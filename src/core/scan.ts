/**
 * Name scanner: builds a lookup of catalog spellings and finds their
 * occurrences in running text, including Turkish case endings
 * ("Gazzâlî'nin") and multi-word names ("Fahreddin er-Râzî").
 * Ported from the VBA 0.3.41 `Vfb*` procedures.
 */
import { MAX_NAME_LENGTH } from "./alias.js";
import type { Person } from "./catalog.js";
import { personKey } from "./catalog.js";
import { excludeMention, hasSuffix, isSentenceInitial, isWorkUsage, schoolRoot } from "./context.js";
import { EPITHETS, EPITHET_KEYS } from "./epithets.js";
import { EPONYMS, fuseLinks, headNisbaAliases, surnameAliases, honorificAliases, nameQuery, nameScore, prominence } from "./names.js";
import { collapseLetters, isSeparator, nameKey, normalize } from "./normalize.js";
import { vbLower, vbTrim } from "./text.js";
import { isValidAlias } from "./alias.js";
import { mid, upper } from "./vb.js";

/** A token: lookup key, 0-based start and length in the source text. */
export interface Token {
  readonly key: string;
  readonly start: number;
  readonly length: number;
}

/** `VfbBreak`: token separator. Hyphens and apostrophes stay inside names. */
export function isBreak(ch: string): boolean {
  if (["-", "'", "’", "‘", "ʻ", "ʼ", "ʿ"].includes(ch)) return false;
  if (ch.length === 0) return true;
  let code = ch.charCodeAt(0);
  if (code >= 65 && code <= 90) code += 32;
  return isSeparator(code);
}

const QUOTES = new Set(["'", "’", "‘", "ʻ", "ʼ", "ʿ", "`", "´"]);

/** `VfbTokens`. */
export function tokenize(text: string, cache: Map<string, string> = new Map()): Token[] {
  const result: Token[] = [];
  let first = -1;
  for (let i = 0; i <= text.length; i++) {
    const ch = i < text.length ? (text[i] as string) : " ";
    if (isBreak(ch)) {
      if (first >= 0) {
        // Quotation marks around a word are not part of it: ‘Gazzâlî’, 'Râzî'.
        let last = i;
        while (first < last && QUOTES.has(text[first] as string)) first++;
        while (last > first && QUOTES.has(text[last - 1] as string)) last--;
        if (first >= last) {
          first = -1;
          continue;
        }
        const value = text.slice(first, last);
        let key = cache.get(value);
        if (key === undefined) {
          key = nameKey(value);
          cache.set(value, key);
        }
        if (key.length > 0) result.push({ key, start: first, length: last - first });
        first = -1;
      }
    } else if (first < 0) {
      first = i;
    }
  }
  return result;
}

const NOISE = new Set("ibn|ibni|bin|bint|ebu|ebul|el|er|es|et|ez|en|ed|bey|han|sah|seyh|imam|haci|seyyid|sair".split("|"));
const COMMON = new Set(
  ("tarif|tanim|ilim|bilim|isim|ad|bilgi|tarih|eser|kitap|usul|fikih|hukuk|hadis|tefsir|kelam|felsefe|mantik|delil|akil|nazar|mana|anlam|hak|" +
    "hakikat|adalet|hayat|olum|dogum|zaman|yer|islam|kuran|sunet|tasavuf|ahlak|fikir|dusunce|mesele|bahis|bab|fasl|fasil|cilt|sayfa|baski|yayin|tercume|" +
    "serh|hasiye|kavram|terim|kisi|sahis|yazar|arastirma|metin|deneme|konu|once|sonra|gore|icin|ile|ve|veya|yahut|ama|fakat|bu|su|bir|" +
    "iki|uc|her|olan|olarak|oldu|dedi|der|gibi|uzere|hakinda").split("|"),
);
/**
 * Words that a part of a longer catalog name must not stand for on its own:
 * places and peoples ("İbn Irâk" → "Irak", "Molla Arap" → "Arap"), titles
 * ("Has Bey" → "Bey"), and ordinary words ("Hayâtî Ahmed Efendi" → "Hayatı",
 * "Cûzcânî, Minhâc-ı Sirâc" → "Minhâc", "Ehlî-i Şîrâzî" → "Ehl-i").
 * The full name still matches.
 */
const WEAK_PARTS = new Set(
  ("irak|arap|arab|acem|turk|kurt|fars|rum|frenk|mogol|tatar|berberi|hint|hind|cin|misir|sam|bagdat|bagdad|mekke|medine|kudus|halep|basra|kufe|horasan|endulus|yemen|hicaz|anadolu|iran|turan|buhara|semerkant|nisabur|merv|herat|belh|kirman|isfahan|tebriz|konya|kahire|dimask|rey|tus|curcan|harezm|kirim|bosna|sirvan|gilan|taberistan|fergana|kasgar|hokand|hive|bursa|edirne|istanbul|sivas|kayseri|amasya|halic|cezire|musul|kerkuk|tus|nehavend|sistan|kabil|gazne|lahor|delhi|dekken|bengal|" +
    "ingiliz|fransiz|alman|italyan|ispanyol|portekiz|rus|yunan|arnavut|bosnak|sirp|bulgar|macar|leh|ozbek|turkmen|kazak|kirgiz|uygur|azeri|cagatay|kipcak|oguz|sudanli|habes|zenci|kipti|yahudi|israilli|hiristiyan|nasrani|mecusi|alevi|kizilbas|surye|suriye|lubnan|urdun|filistin|sudan|libya|tunus|cezayir|fas|merakes|misr|ezher|ayasofya|asya|avrupa|afrika|amerika|hindistan|pakistan|afganistan|endonezya|malezya|mardin|diyarbakir|erzurum|trabzon|selcuk|kudus|sam|beyrut|kahire|iskenderiye|" +
    "kureys|temim|kays|kinde|tayy|ezd|evs|hazrec|sakif|huzeyl|gatafan|hevazin|benihasim|hasimi|emeviler|abbasiler|eyyubi|merini|resuli|behmeni|abbadi|seybani|pehlevi|kolemen|memluk|tolunogullari|husayni|huseyni|naksi|bahai|suudi|vehhabi|zeydi|ismaili|haricy|" +
    "muderris|fakih|alim|edip|edib|asker|silahdar|hekimbasi|hadim|kadi|nakibulesraf|musahip|defterdar|nisanci|reisulkuttab|kazasker|mutevelli|imamhatip|hattat|musikisinas|neyzen|tabip|tabib|hekim|cerrah|muneccim|mimar|ressam|nakkas|mucellit|katib|sair|ozan|asik|devlet|millet|ummet|seriyye|seriye|" +
    "molla|efendi|bey|pasa|sultan|seyh|hoca|dede|baba|kadi|emir|mir|han|hatun|banu|aga|celebi|mevlana|sah|hace|hafiz|hakim|mufti|imam|mollasi|sehzade|vezir|beylerbeyi|kethuda|reis|nakib|hatip|vaiz|muezin|katip|kadiasker|seyhulislam|" +
    "hayat|hayati|hayatin|has|cem|ehl|ehli|bil|minhac|nur|sir|can|gul|bahar|derd|dert|hal|kal|ask|asik|garip|garib|fakir|dervis|sufi|zahit|zahid|abit|abid|arif|salih|emin|kerim|rahim|aziz|celil|cemal|kemal|celal|hilal|hayal|sabit|baki|fani|kadim|cedid|kebir|sagir|kucuk|buyuk|yeni|eski|kara|ak|kizil|sari|ali|ulu|has|hassa|nesih|talik|sulus|divani|rika|mevlevi|bektasi|nakisbendi|kadiri|rifai|halveti|celveti|hanefi|safii|maliki|hanbeli|sunni|sii|abbasi|emevi|fatimi|selcuklu|osmanli|ilhanli|timurlu|safevi|babur").split("|").map(collapseLetters),
);

/**
 * Ordinary words that open a sentence right before a name: kinship and
 * teacher–student nouns, sentence adverbs ("Hocası Cüveynî", "Nitekim Gazzâlî").
 */
const OPENERS =
  /^(hoca|talebe|ogrenci|baba|ogul|oglu|kizi|kardes|agabey|dede|anne|esi|hanimi|seyhi|muridi|halife|cagdas|muasir|rakib|rakip|sarih|hami|veziri|sultani|hukumdar|muasiri|ustad|ustat|nitekim|ayrica|ozellikle|hatta|yine|mesela|ornegin|dolayisiyla|boylece|nihayet|ilk|son|gene|bilhasa|bunun|bunlardan|onun|ondan|sonraki|onceki|daha|ancak|fakat|lakin|zira|cunku|hem|ne|ya|ise|keza|kezalik)/;

const NAME_LINKS = new Set("b|ibn|ibni|bin|bint|ebu|ebul|el|er|es|et|ez|en|ed".split("|"));

export const isNoise = (key: string): boolean => NOISE.has(key);
export const isCommonWord = (key: string): boolean => COMMON.has(key);
export const isNameLink = (key: string): boolean => NAME_LINKS.has(key);

/** `VfbWordCount`. */
export function wordCount(value: string): number {
  let count = 0;
  let inWord = false;
  for (const ch of value) {
    if (isBreak(ch)) inWord = false;
    else if (!inWord) {
      count++;
      inWord = true;
    }
  }
  return count;
}

/** `VfbInitialUpper`: the first cased letter is a capital (uncased scripts count as capital). */
export function initialUpper(value: string): boolean {
  if (["el-", "er-", "es-", "et-", "ez-", "en-", "ed-", "eş-"].includes(vbLower(value.slice(0, 3)))) value = value.slice(3);
  for (const ch of value) {
    if (upper(ch) !== vbLower(ch)) return ch === upper(ch);
    const code = ch.charCodeAt(0);
    // Scripts without case (Arabic, CJK …) count as capitalised; emoji and symbols do not.
    if (((code >= 0x590 && code <= 0x1fff) || code >= 0x2e80) && /\p{L}/u.test(ch)) return true;
  }
  return false;
}

/** `VfbWithoutArticles`: normalized name without Arabic article tokens. */
export function withoutArticles(value: string): string {
  let result = ` ${normalize(value)} `;
  for (const article of ["el", "er", "es", "et", "ez", "en", "ed"]) result = result.split(` ${article} `).join(" ");
  return vbTrim(result);
}

/** `VfbNaturalOrder`: "Râzî, Fahreddin" -> "Fahreddin Râzî". */
export function naturalOrder(value: string): string {
  const comma = value.indexOf(",");
  if (comma < 0 || value.indexOf(",", comma + 1) >= 0) return "";
  const first = vbTrim(value.slice(comma + 1));
  const last = vbTrim(value.slice(0, comma));
  return first.length > 0 && last.length > 0 ? `${first} ${last}` : "";
}

function compoundAlias(value: string): boolean {
  if (!value.includes("-")) return false;
  if (!value.includes("'") && !value.includes("’")) return false;
  const key = nameKey(value);
  return !(key.startsWith("ebu") || key.startsWith("abu"));
}

function independentAlias(value: string, shortName: string): boolean {
  const key = nameKey(value);
  const root = schoolRoot(key);
  // A school word alone names only its eponym ("Eş‘arî" → Ebü’l-Hasan el-Eş‘arî); the context rules tell adjective from person.
  return root.length > 0 ? key === nameKey(shortName) || (key === root && EPONYMS[root] === shortName) : true;
}

/** Lookup of catalog spellings (`VfbBuildPatterns`, `VfbIndexRow`, `VfbAddPattern`). */
export class NameIndex {
  /** Spelling key -> catalog row indices. */
  readonly patterns = new Map<string, Set<number>>();
  /** First three letters of every key, to stop scanning early. */
  readonly heads = new Set<string>();
  /** Spelling key -> word counts it was seen with (strict mode). */
  readonly shapes = new Map<string, Set<number>>();
  readonly cache = new Map<string, string>();

  constructor(readonly people: readonly Person[]) {
    people.forEach((person, i) => this.indexRow(person, i));
  }

  private add(value: string, index: number, epithet = false): void {
    let key = this.cache.get(value);
    if (key === undefined) {
      key = nameKey(value);
      this.cache.set(value, key);
    }
    if (key.length < 3 || value.length > MAX_NAME_LENGTH || value === "---") return;
    // An epithet ("Mevlânâ", "Resûlullah") is a name although the word is a title elsewhere.
    if (!epithet && (isNoise(key) || isCommonWord(key))) return;
    let ids = this.patterns.get(key);
    if (!ids) {
      ids = new Set();
      this.patterns.set(key, ids);
      this.heads.add(key.slice(0, 3));
    }
    ids.add(index);
    let counts = this.shapes.get(key);
    if (!counts) {
      counts = new Set();
      this.shapes.set(key, counts);
    }
    counts.add(wordCount(value));
    // "A‘lem": an ayn inside a word does not split it in running text.
    counts.add(wordCount(value.replace(/(\p{L})[‘’'ʿʾ`](\p{L})/gu, "$1$2")));
  }

  /** A single word of a longer name that is a place, title or ordinary word. */
  private weakPart(part: string, whole: string): boolean {
    if (wordCount(whole) < 2) return false;
    const plain = collapseLetters(nameKey(part));
    // "Bursalı", "Tunuslu": a place adjective is not the person's name on its own; "-iyye" abstracts too.
    if (plain.length > 5 && /(li|lu)$/.test(plain) && /[lL][ıiuü]$/u.test(part)) return true;
    if (/iye$|iyye$/.test(plain)) return true;
    // "İshak Hocası" → "Hocası": kinship and teacher words are not names on their own.
    if (OPENERS.test(plain)) return true;
    // "Minhâc-ı Sirâc" → "Minhâc-ı": the izafet ending does not make it a name.
    return WEAK_PARTS.has(collapseLetters(nameKey(part))) || WEAK_PARTS.has(collapseLetters(nameKey(part.replace(/[-‐][ıiuü]$/u, ""))));
  }

  private indexRow(person: Person, i: number): void {
    for (const epithet of EPITHETS[person.shortName] ?? []) this.add(epithet, i, true);
    // "Ahund el-Horasânî": a word of the short name with a nisba of the full name.
    for (const alias of headNisbaAliases(person.shortName, person.fullName)) this.add(alias, i);
    for (const alias of surnameAliases(person.shortName, person.source)) this.add(alias, i);
    const values = [person.shortName, person.fullName, person.originalName];
    values.forEach((value, j) => {
      this.add(value, i);
      this.add(withoutArticles(value), i);
      // "Ebû Ca‘fer Nüfeylî": without the article, an ayn inside a word keeps it one word.
      if (/\p{L}[‘’'ʿʾ`]\p{L}/u.test(value)) this.add(withoutArticles(value.replace(/(\p{L})[‘’'ʿʾ`](\p{L})/gu, "$1$2")), i);
      // "Ebü’l-Abbâs Seffâh": only the separate articles go, the künye keeps its own.
      const bare = value.replace(/(^|\s)(el|er|es|et|ez|en|ed|eş)[-‐](?=\p{L})/giu, "$1");
      if (bare !== value) this.add(bare, i);
      if (j === 0) {
        const natural = naturalOrder(value);
        if (natural.length > 0) {
          this.add(natural, i);
          this.add(withoutArticles(natural), i);
        }
      }
      for (const token of tokenize(value, this.cache)) {
        const alias = value.slice(token.start, token.start + token.length);
        // Short catalog names supply independent aliases; components of long
        // names do not, except joined forms such as "Kıvâmü's-sünne".
        if ((j === 0 || compoundAlias(alias)) && independentAlias(alias, person.shortName) && !this.weakPart(alias, value)) this.add(alias, i);
      }
      if (j === 1) {
        // Künye + nisba, the usual short form of a long name:
        // "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî" → "Ebû Hâmid el-Gazzâlî".
        // Honorifics before it ("Hüccetü’l-İslâm") and a second nisba after it ("et-Tûsî") are allowed.
        const words = value.trim().split(/\s+/);
        const at = words.findIndex((w, k) => k < 4 && /^(Eb[ûüu]|Ebi|Ümm[üu]?)([’'‘]|$|-)/i.test(w));
        if (at >= 0 && words.length > at + 3) {
          const head = words[at] as string;
          const kunya = /[’'‘-]/.test(head) ? head : `${head} ${words[at + 1] ?? ""}`;
          const nisbas = words.slice(at + 2).filter((w) => /^(el|er|es|et|ez|en|ed|eş|ed)[-‐]\p{Lu}/iu.test(w));
          for (let n = 0; n < Math.min(nisbas.length, 3); n++) {
            const alias = `${kunya} ${nisbas[n]}`;
            this.add(alias, i);
            this.add(withoutArticles(alias), i);
          }
          if (nisbas.length >= 2) this.add(`${kunya} ${nisbas[0]} ${nisbas[1]}`, i);
        }
        // Lakap + nisba, the other usual short form: "Fahrü’l-İslâm Alî b. Muhammed … el-Pezdevî"
        // → "Fahrü’l-İslâm el-Pezdevî", also written joined: "Fahrülislâm Pezdevî".
        for (const alias of honorificAliases(value)) {
          // Joined or apart, lakap and künye each: "Fahrülislâm Ebü’l-Usr", "Fahrü’l-İslâm Ebülusr" …
          const [first, ...rest] = alias.split(" ");
          const forms = new Set([alias, fuseLinks(alias), [fuseLinks(first as string), ...rest].join(" "), [first, ...rest.map(fuseLinks)].join(" ")]);
          for (const form of forms) {
            this.add(form, i);
            this.add(withoutArticles(form), i);
          }
        }
      }
      if (j === 0) {
        // An apostrophe inside a word does not split it: "Di‘bil" must not give "bil".
        for (const part of normalize(value.replace(/(\p{L})['’‘ʻʼʿ](\p{L})/gu, "$1$2")).split(" ")) {
          if (independentAlias(part, person.shortName) && !this.weakPart(part, value)) this.add(part, i);
        }
      }
    });
  }

  has(key: string): boolean {
    return this.patterns.has(key);
  }

  shapeAllowed(key: string, count: number): boolean {
    return this.shapes.get(key)?.has(count) ?? false;
  }

  /** `VfbRankRows`: catalog rows for a matched key, best first, one per person. */
  rank(key: string, query: string): Person[] {
    const ids = this.patterns.get(key);
    if (!ids) return [];
    const seen = new Set<string>();
    const ranked: Array<{ rank: number; person: Person }> = [];
    for (const index of ids) {
      const person = this.people[index] as Person;
      const id = personKey(person);
      if (seen.has(id)) continue;
      seen.add(id);
      const best = Math.max(nameScore(person.shortName, query), nameScore(person.fullName, query), nameScore(person.originalName, query));
      const rank = best * 1000 + prominence(person.shortName, nameKey(query));
      let position = ranked.length;
      while (position > 0 && (ranked[position - 1] as { rank: number }).rank < rank) position--;
      ranked.splice(position, 0, { rank, person });
    }
    return ranked.map((entry) => entry.person);
  }
}

/** `VfbJoinable`: tokens of one name are separated by spaces only ("b." "i." "v." allowed). */
function joinable(gap: string, previous: string): boolean {
  // An ayn or hamza ending the previous word ("Ebû Mutî‘ el-Belhî", "Akra‘ b. Hâbis") is part of that word.
  gap = gap.replace(/^[‘’'ʿʾ`](?=[ \t\u00a0])/u, "");
  for (const ch of gap) {
    if (ch === "." && (previous === "b" || previous === "i" || previous === "v")) continue;
    if (ch !== " " && ch !== "\t" && ch !== " ") return false;
  }
  return true;
}

function gapBetween(text: string, previous: Token, token: Token): string {
  return text.slice(previous.start + previous.length, token.start);
}

/** `VfbNameSpanEnd`: last token index of the capitalised name run starting at `first`. */
/**
 * Capitalised words of peoples, faiths and regions that follow a name without
 * being part of it: "Ziya Gökalp Türk sosyolojisini", "Gazzâlî İslâm
 * düşüncesinde", "Fârâbî Yunan felsefesini". They end a run of capitalised words.
 */
const SPAN_STOP = new Set(
  ("turk|turkiye|turkce|turkcu|islam|islami|arap|arapca|fars|farsca|osmanli|osmanlica|selcuklu|bati|batili|dogu|dogulu|avrupa|avrupali|yunan|yunanca|roma|latin|latince|" +
    "hint|cin|alman|almanca|ingiliz|ingilizce|fransiz|fransizca|rus|rusca|amerikan|amerika|hiristiyan|hiristiyanlik|yahudi|yahudilik|sunni|sii|emevi|abbasi|endulus|iran|misir|anadolu|kuran|cumhuriyet|tanzimat|mesrutiyet")
    .split("|")
    .map(collapseLetters),
);

function nameSpanEnd(text: string, tokens: readonly Token[], first: number, index: NameIndex): number {
  let end = first;
  for (let i = first; i < tokens.length; i++) {
    const token = tokens[i] as Token;
    const raw = text.slice(token.start, token.start + token.length);
    let key = token.key;
    if (hasSuffix(raw)) key = nameKey(nameQuery(raw));
    if (isCommonWord(key)) break;
    if (isNoise(key) && !isNameLink(key)) break;
    // A lower-case word ("devlet", "medrese") never continues a name, catalog key or not.
    if (!isNameLink(key) && !initialUpper(raw)) break;
    if (i > first && !joinable(gapBetween(text, tokens[i - 1] as Token, token), (tokens[i - 1] as Token).key)) break;
    if (i > first && SPAN_STOP.has(collapseLetters(key))) break;
    // A title with an article after the name, inflected or called a work: "Şâfiî er-Risâle'yi",
    // "Gazzâlî el-Mustasfâ'da", "Gazzâlî el-Mustasfâ adlı eserinde". The name stops before it.
    if (i > first && /^(el|er|es|et|ez|en|ed|eş)[-‐]/i.test(raw) && (hasSuffix(raw) || isWorkUsage(text.slice(token.start + token.length, token.start + token.length + 40)))) break;
    if (!isNameLink(key)) end = i;
    if (hasSuffix(raw)) break;
  }
  return end;
}

/** A name occurrence found in text. */
export interface Match {
  /** 0-based start in the scanned text. */
  readonly start: number;
  readonly length: number;
  /** Literal text, including a Turkish ending. */
  readonly text: string;
  /** Matched spelling key. */
  readonly key: string;
}

/**
 * `VfbTextMatches`: longest catalog spellings in `text`. In `strict` mode
 * (person discovery) the whole capitalised name run must match, single words
 * must be capitalised and not common words, and context rules (school
 * adjectives, terms, citations) exclude non-person uses.
 */
export function findMatches(text: string, index: NameIndex, strict = false, citationCheck = true): Match[] {
  const result: Match[] = [];
  const tokens = tokenize(text, index.cache);
  let i = 0;
  while (i < tokens.length) {
    const firstToken = tokens[i] as Token;
    const first = firstToken.start;
    let key = "";
    let plainKey = ""; // the same without "el-/er-…": "Ebû Bekir er-Râzî" ~ "Ebû Bekir Râzî"
    let bestEnd = 0;
    let bestIndex = i;
    let bestKey = "";
    const spanEnd = strict ? nameSpanEnd(text, tokens, i, index) : i;
    for (let j = i; j < tokens.length; j++) {
      const token = tokens[j] as Token;
      const finish = token.start + token.length;
      if (finish - first > MAX_NAME_LENGTH) break;
      if (j > i && !joinable(gapBetween(text, tokens[j - 1] as Token, token), (tokens[j - 1] as Token).key)) break;
      let tokenKey = token.key;
      const raw = text.slice(token.start, finish);
      // In strict mode a lower-case word ends the name ("Cem sultan"), links ("b.", "el-") aside.
      if (strict && j > i && !isNameLink(token.key) && !initialUpper(raw)) break;
      let inflected = false;
      // A Turkish ending after an apostrophe is never part of the name: "Minhac'ı" is not "Minhâcî".
      if (hasSuffix(raw) && (strict || !index.has(collapseLetters(key + tokenKey)))) {
        tokenKey = nameKey(nameQuery(raw));
        inflected = true;
      }
      key = collapseLetters(key + tokenKey);
      const article = /^(el|er|es|et|ez|en|ed|eş)[-‐]/i.exec(raw);
      plainKey = collapseLetters(plainKey + (article && j > i ? nameKey(nameQuery(raw.slice(article[0].length))) : tokenKey));
      if (key.length >= 3 && !index.heads.has(key.slice(0, 3)) && !index.heads.has(plainKey.slice(0, 3))) break;
      if (index.has(key)) {
        bestEnd = finish;
        bestIndex = j;
        bestKey = key;
      } else if (plainKey !== key && index.has(plainKey)) {
        bestEnd = finish;
        bestIndex = j;
        bestKey = plainKey;
      }
      if (inflected) break;
    }
    if (bestEnd > first) {
      let accepted = true;
      if (strict) {
        if (!index.shapeAllowed(bestKey, bestIndex - i + 1)) accepted = false;
        if (spanEnd > bestIndex) accepted = false;
        if (bestIndex === i) {
          const raw = mid(text, first + 1, bestEnd - first);
          const own = nameKey(nameQuery(raw));
          if ((isCommonWord(own) && !EPITHET_KEYS.has(own)) || !initialUpper(raw)) accepted = false;
        }
        if (accepted) accepted = !excludeMention(text, first, bestEnd - first, bestKey, citationCheck);
      }
      if (accepted) result.push({ start: first, length: bestEnd - first, text: text.slice(first, bestEnd), key: bestKey });
      i = Math.max(bestIndex, spanEnd) + 1;
    } else if (strict && spanEnd > i && !index.has(firstToken.key) && OPENERS.test(firstToken.key) && isSentenceInitial(text, first)) {
      // "Hocası Cüveynî …": a capitalised ordinary word opening the sentence
      // is not the start of the name run; look again from the next word.
      i++;
    } else {
      i = spanEnd + 1;
    }
  }
  return result;
}

// ---- People absent from the catalog (VBA VfjUnknownMatches) ----

const ROLES = new Set(["dr", "prof", "doc", "yazar", "muellif", "sair", "filozof", "dusunur", "tarihci"]);
const ORGANIZATIONS = "universite|fakulte|enstitu|yayin|matbaa|kutuphane|devleti|bakanlik|belediye|vakfi|dernegi|konseyi|ansiklopedisi|sempozyum|kongresi".split("|");

const isCased = (text: string): boolean => text.toUpperCase() !== text.toLowerCase();

function roleGap(gap: string, role: string): boolean {
  if (gap.length > 8) return false;
  for (const ch of gap) {
    if (ch === " " || ch === "\t" || ch === ":") continue;
    if (ch !== "." || (role !== "dr" && role !== "prof" && role !== "doc")) return false;
  }
  return true;
}

/**
 * Capitalised names right after a person role ("Dr.", "müellif", "şair" …)
 * that are not in the catalog, used by Şahıs dökümü. Organisations, common
 * words and context exclusions are rejected.
 */
export function findUnknownPeople(text: string, index: NameIndex): Match[] {
  const result: Match[] = [];
  const tokens = tokenize(text, index.cache);
  let i = 0;
  while (i < tokens.length) {
    const token = tokens[i] as Token;
    let raw = text.slice(token.start, token.start + token.length);
    if (ROLES.has(token.key) || !isCased(raw) || !initialUpper(raw) || isCommonWord(token.key)) {
      i++;
      continue;
    }
    const last = nameSpanEnd(text, tokens, i, index);
    let strong = 0;
    let reject = false;
    let role = false;
    if (i > 0) {
      const previous = tokens[i - 1] as Token;
      role = ROLES.has(previous.key) && roleGap(gapBetween(text, previous, token), previous.key);
    }
    for (let j = i; j <= last; j++) {
      const part = tokens[j] as Token;
      raw = text.slice(part.start, part.start + part.length);
      if (ORGANIZATIONS.some((word) => part.key.includes(word)) || isCommonWord(part.key)) reject = true;
      if (!isNameLink(part.key) && !isNoise(part.key)) {
        if (isCased(raw) && initialUpper(raw)) strong++;
        else reject = true;
      }
    }
    const first = token.start;
    const end = tokens[last] as Token;
    const finish = end.start + end.length;
    raw = text.slice(first, finish);
    if (!reject && role && strong >= 1 && isValidAlias(raw) && !excludeMention(text, first, finish - first, "", true)) {
      result.push({ start: first, length: finish - first, text: raw, key: "" });
    }
    i = last + 1;
  }
  return result;
}
