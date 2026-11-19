/**
 * Task pane decisions that do not need Word. Word calls live in word.ts so
 * that everything here is unit tested.
 */
import {
  type Catalog,
  type Person,
  type PrefixKey,
  alreadyHasDate,
  confirmationCandidates,
  formatDate,
  isValidAlias,
  isValidDate,
  nameCandidates,
  nameQuery,
  normalize,
  personKey,
  safeCandidate,
  startsWithDate,
  strength,
  textName,
} from "../core/index.js";

export interface Lookup {
  /** Selected name without a Turkish ending. */
  readonly query: string;
  /** Strong candidates (same name), then partial ones. */
  readonly candidates: readonly Person[];
  readonly strong: ReadonlySet<string>;
  /** Best name score per person key: 4/3 same name, 2 part of the name, 1 similar spelling. */
  readonly scores: ReadonlyMap<string, number>;
  /** Set when exactly one strong candidate exists: no choice needed. */
  readonly automatic: Person | null;
  /** A title before the name that the lookup left out ("Hz. ", "İmam "), or "". */
  readonly title?: string;
}

/** Titles written before a name: "Hz. Hatice", "Hazret-i Ömer", "İmam Şâfiî", "Şeyh Şâzelî". */
const TITLE = /^(hz\.|hz|hazret-?i|hazreti|hazret|imam|imâm|şeyh|şeyhülislâm|şeyhülislam|mevlânâ|mevlana|hâce|hace|seyyid|şerif|hâfız|hafız)(?:\s+|(?<=\.)(?=\p{L}))/iu;

/** The title before a selected name and the name itself: "Hz. Hatice" → ["Hz. ", "Hatice"]. */
export function splitTitle(alias: string): [string, string] {
  // "İ" does not match "i" case-insensitively in JavaScript; the same length "I" does.
  const match = TITLE.exec(alias.replace(/^İ/, "I"));
  if (!match) return ["", alias];
  const rest = alias.slice(match[0].length);
  return rest.trim().length >= 2 ? [alias.slice(0, match[0].length), rest] : ["", alias];
}

/**
 * Candidates for a selected name, as the VBA "Ölüm tarihi ekle" command
 * gathers them: name matches only (no description matches), one entry per
 * person, strong matches first.
 */
export function lookup(catalog: Catalog, alias: string): Lookup {
  const first = lookupName(catalog, alias);
  const [title, rest] = splitTitle(alias);
  if (!title) return first;
  // A title written before the name ("Hz. Hatice", "İmam Şâfiî"); when it is part of
  // the catalog name ("Molla Hüsrev", "Kadı Abdülcebbâr") the name as written wins.
  if (first.strong.size > 0) return first;
  const bare = lookupName(catalog, rest);
  return bare.candidates.length > 0 ? { ...bare, title } : first;
}

function lookupName(catalog: Catalog, alias: string): Lookup {
  const found = confirmationCandidates(catalog.find(alias, false), alias);
  const { exact, partial } = nameCandidates(found, alias);
  const strong = new Set(exact.map(personKey));
  return {
    query: nameQuery(alias),
    candidates: [...exact, ...partial],
    strong,
    scores: new Map(found.map((person) => [personKey(person), strength(person, alias)])),
    automatic: safeCandidate(found, alias),
  };
}

export type DatePlan =
  | { readonly ok: true; readonly text: string }
  | { readonly ok: false; readonly reason: string };

/**
 * Whether a date may be written after the name, and its text. `following` is
 * the rest of the paragraph after the selected name.
 */
export function planDate(person: Person, prefix: PrefixKey, following: string): DatePlan {
  if (!isValidDate(person.deathDate)) {
    return { ok: false, reason: "Katalogda bu şahıs için kullanılabilir bir ölüm tarihi yok; tarih eklenmedi." };
  }
  if (startsWithDate(following) || alreadyHasDate(following, person.deathDate)) {
    return { ok: false, reason: "Adın hemen arkasında zaten bir tarih var; mevcut tarihe dokunulmadı." };
  }
  return { ok: true, text: formatDate(person.deathDate, prefix) };
}

/**
 * The date written directly after a name ("(ö. 505/1111)", "(1913-1998)"),
 * as it stands in the text, or null. Only a parenthesis that startsWithDate
 * accepts, ending at its own ")" on the same line.
 */
export function leadingDate(following: string): string | null {
  if (!startsWithDate(following)) return null;
  const open = following.indexOf("(");
  const close = following.indexOf(")", open);
  if (open < 0 || close < 0) return null;
  const date = following.slice(open, close + 1);
  return /[\r\n\u0007\u000b]/.test(date) || date.length > 60 ? null : date;
}

/** Whether two date texts are the same apart from spacing and case. */
export function sameDate(a: string, b: string): boolean {
  const key = (t: string): string => t.replace(/\s+/g, "").toLocaleLowerCase("tr");
  return key(a) === key(b);
}

/**
 * Literal spellings searched before the selection to warn about an earlier
 * mention: the selected spelling, the three catalog names, their reading
 * order and the family part of the short name ("Râzî" of "Fahreddin er-Râzî").
 * Very short or invalid values are skipped; at most 12 spellings.
 */
export function earlierSpellings(person: Person, alias: string): string[] {
  const result: string[] = [];
  const seen = new Set<string>();
  const add = (value: string): void => {
    value = value.trim();
    const key = normalize(value);
    if (value.length < 3 || value.length > 200 || key.length < 3 || !isValidAlias(value) || seen.has(value)) return;
    seen.add(value);
    result.push(value);
  };
  add(nameQuery(alias));
  for (const name of [person.shortName, person.fullName, person.originalName]) {
    add(name.includes(",") ? textName(name) : name);
  }
  const short = person.shortName.includes(",")
    ? person.shortName.slice(0, person.shortName.indexOf(","))
    : (person.shortName.trim().split(/[\s\-–]+/).pop() ?? "");
  add(short);
  return result.slice(0, 12);
}

/** Content control tag of a Vefeyât date: short, stable per person. */
export function dateTag(person: Person): string {
  return `vefeyat:date:${hash(personKey(person))}`;
}

function hash(value: string): string {
  // FNV-1a 32-bit; only used to group dates of one person, not for security.
  let h = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    h ^= value.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

/** Short note explaining a non-exact candidate, or "" for the same name. */
export function matchHint(score: number): string {
  if (score >= 3) return "";
  if (score === 2) return "adın bir parçası";
  return "benzer yazım, doğrulayın";
}

/** One-line summary of a candidate for the list. */
export function candidateLabel(person: Person): string {
  const name = textName(person.shortName);
  return person.deathDate ? `${name} ${person.deathDate}` : name;
}

export type View = "secili" | "dokum" | "teslim" | "katalog" | "yardim" | "ara" | "saglik" | "kaynak" | "gecmis" | "toplu" | "kimlik" | "istisna" | "kapsam" | "tercihler";
export type Tool = "tutarlilik" | "rapor" | "dizin" | "dizinkaldir" | "dizinsil" | "temizlik";

export interface Launch {
  readonly view: View;
  readonly action: "insert" | "spelling" | "dia" | "audit" | "resume" | "prefer" | "clearSpelling" | "shading" | "noproof" | null;
  readonly tool: Tool | null;
}


/**
 * Which view and action a ribbon or context-menu button asked for, from the
 * task pane URL query (?action=insert, ?view=dokum, ?tool=dizin …).
 * Unknown values fall back to the selected-name view.
 */
export function parseLaunch(search: string): Launch {
  const query = new URLSearchParams(search);
  const action = query.get("action");
  const tool = query.get("tool");
  const view = query.get("view");
  if (action === "insert" || action === "spelling" || action === "dia" || action === "prefer") return { view: "secili", action, tool: null };
  if (action === "clearSpelling" || action === "shading" || action === "noproof") return { view: "tercihler", action, tool: null };
  if (tool === "tutarlilik" || tool === "rapor" || tool === "dizin" || tool === "dizinkaldir" || tool === "dizinsil" || tool === "temizlik") {
    return { view: "teslim", action: null, tool };
  }
  if (view === "denetim") return { view: "secili", action: "audit", tool: null };
  if (view === "devam") return { view: "secili", action: "resume", tool: null };
  if (view === "dokum") return { view: "dokum", action: null, tool: null };
  if (view === "teslim" || view === "katalog" || view === "yardim" || view === "ara" || view === "saglik" || view === "kaynak" || view === "gecmis" || view === "toplu" || view === "kimlik" || view === "istisna" || view === "kapsam" || view === "tercihler") return { view, action: null, tool: null };
  if (view === "ayarlar") return { view: "katalog", action: null, tool: null };
  if (view === "kisayol") return { view: "tercihler", action: null, tool: null };
  return { view: "secili", action: null, tool: null };
}

/**
 * Why the ribbon's "Ölüm tarihi ekle" cannot write at once, or null when the
 * selection is unambiguous and the normal checks may run.
 */
export function quickInsertBlocker(hasCatalog: boolean, alias: string | null, found: Lookup | null): string | null {
  if (!hasCatalog) return "Katalog yüklenemedi. Vefeyât → Katalog düğmesinden CSV dosyasını seçin.";
  if (!alias) return "Metinde bir şahıs adı seçip yeniden deneyin.";
  if (!found || found.candidates.length === 0) return "Katalogda bu adla eşleşen şahıs yok.";
  if (!found.automatic) return "Birden fazla aday veya benzer yazım var. Listeden şahsı seçip Ölüm tarihi ekle'ye basın.";
  return null;
}

// ---- Document settings (VBA: VfpScope, VfpPrefix, VfuNameForm, VfpAuditMode) ----

export type Scope = "first" | "chapter" | "all" | "none";
export type NameForm = "ask" | "near" | "short" | "full" | "original";
export type AuditMode = "all" | "undated";

export interface Settings {
  readonly prefix: PrefixKey;
  readonly scope: Scope;
  readonly nameForm: NameForm;
  readonly auditMode: AuditMode;
}

export const DEFAULT_SETTINGS: Settings = { prefix: "o", scope: "first", nameForm: "ask", auditMode: "all" };

export const SETTING_CHOICES = {
  prefix: { original: "Kaynakta olduğu gibi", o: "ö.", ol: "öl.", v: "v.", d: "d. (İngilizce)" },
  scope: { first: "Yalnız ilk kullanım", chapter: "Her bölümde ilk kullanım", all: "Bütün kullanımlar", none: "Yalnız denetle" },
  nameForm: { ask: "Her seferinde sor", near: "Seçili ada yakın", short: "Kısa ad", full: "Tam / uzun ad", original: "Orijinal ad" },
  auditMode: { all: "Tüm geçişler", undated: "Tarihi olmayanlar" },
} as const;

export type SettingName = keyof typeof SETTING_CHOICES;

/** A stored value, or the default when missing or unknown. */
export function settingValue<K extends SettingName>(name: K, stored: unknown): Settings[K] {
  const choices = SETTING_CHOICES[name] as Record<string, string>;
  return (typeof stored === "string" && stored in choices ? stored : DEFAULT_SETTINGS[name]) as Settings[K];
}

/** One-line summary shown at the top of the pane. */
export function settingsSummary(settings: Settings): string {
  return [
    `Kısaltma: ${SETTING_CHOICES.prefix[settings.prefix]}`,
    SETTING_CHOICES.scope[settings.scope],
    `Ad biçimi: ${SETTING_CHOICES.nameForm[settings.nameForm]}`,
  ].join(" · ");
}

/**
 * Whether the scope setting blocks a date or requires the earlier-mention
 * check. "none" never writes; "all" writes at every mention.
 */
export function scopeRule(scope: Scope): { allowed: boolean; checkEarlier: boolean; reason?: string } {
  if (scope === "none") {
    return { allowed: false, checkEarlier: false, reason: "Tarih ekleme ayarı \"Yalnız denetle\": tarih eklenmez. Vefeyât → Tarih ekleme menüsünden değiştirebilirsiniz." };
  }
  return { allowed: true, checkEarlier: scope === "first" || scope === "chapter" };
}

const FORM_LABELS: Record<Exclude<NameForm, "ask">, string> = {
  near: "Seçili ada yakın yazım",
  short: "Kısa ad",
  full: "Tam / uzun ad",
  original: "Orijinal ad",
};

/** Puts the preferred spelling (Ad biçimi) first; "ask" keeps the closest-first order. */
export function orderByNameForm<T extends { label: string }>(options: readonly T[], form: NameForm): T[] {
  if (form === "ask") return [...options];
  const wanted = FORM_LABELS[form];
  return [...options.filter((o) => o.label === wanted), ...options.filter((o) => o.label !== wanted)];
}

/**
 * Which inventory rows "Tüm belgeyi denetle" visits (VBA `VfwPrepare`): with
 * "Yalnız ilk kullanım", verified repeats are skipped; with "Tarihi
 * olmayanlar", mentions that already carry a date are skipped.
 */
export function auditRows<T extends { person: unknown; first: boolean; firstInChapter?: boolean; date: string; dateCheck?: string | null }>(rows: readonly T[], settings: Settings): T[] {
  return rows.filter((row) => {
    // A wrong written date is always worth a stop, repeat or not.
    if (row.dateCheck === "different" || row.dateCheck === "inconsistent") return true;
    if (settings.scope === "first" && row.person !== null && !row.first) return false;
    if (settings.scope === "chapter" && row.person !== null && !(row.firstInChapter ?? row.first)) return false;
    if (settings.auditMode === "undated" && row.date) return false;
    return true;
  });
}

// ---- Identity confirmations (VBA: Vdi, contextual identity) ----

/** Most confirmations kept per document (VBA allowed 200 contextual + registry). */
export const CONFIRMATION_LIMIT = 2000;

/**
 * Key of one mention: the literal name with the normalized words just before
 * and after it in the same paragraph. Edits elsewhere keep it; rewriting the
 * sentence around the name invalidates it, as in the VBA signature.
 */
export function mentionKey(literal: string, before: string, after: string): string {
  // A date written right after the name (often by this add-in) is not context.
  const trimmed = after.replace(/^[ \u00a0]+/, "");
  if (startsWithDate(trimmed)) {
    const closing = trimmed.indexOf(")");
    if (closing >= 0 && closing < 80) after = trimmed.slice(closing + 1);
  }
  const head = normalize(before).split(" ").filter(Boolean).slice(-6).join(" ");
  const tail = normalize(after).split(" ").filter(Boolean).slice(0, 6).join(" ");
  return `${head}|${literal}|${tail}`;
}

export type Confirmations = Readonly<Record<string, string>>;

/** Parses the stored confirmations (JSON object mention key -> person key). */
export function parseConfirmations(stored: unknown): Record<string, string> {
  if (typeof stored !== "string") return {};
  try {
    const value: unknown = JSON.parse(stored);
    if (!value || typeof value !== "object" || Array.isArray(value)) return {};
    return Object.fromEntries(Object.entries(value as Record<string, unknown>).filter((e): e is [string, string] => typeof e[1] === "string"));
  } catch {
    return {};
  }
}

/** Adds one confirmation; the oldest are dropped beyond the limit. */
export function addConfirmation(map: Confirmations, key: string, person: string): Record<string, string> {
  const next: Record<string, string> = { ...map };
  delete next[key];
  next[key] = person;
  const keys = Object.keys(next);
  for (const old of keys.slice(0, Math.max(0, keys.length - CONFIRMATION_LIMIT))) delete next[old];
  return next;
}

/** The confirmed candidate for a mention, if it is still among the candidates. */
export function confirmedPerson<T extends Person>(map: Confirmations, key: string, candidates: readonly T[]): T | null {
  const wanted = map[key];
  return wanted === undefined ? null : candidates.find((p) => personKey(p) === wanted) ?? null;
}

// ---- Belge istisnaları ("Şahıs değil") ----

/** Key of an excluded phrase: the normalized name without its Turkish ending. */
export function exclusionKey(literal: string): string {
  return normalize(nameQuery(literal));
}

export function parseExclusions(stored: unknown): string[] {
  if (typeof stored !== "string") return [];
  try {
    const value: unknown = JSON.parse(stored);
    return Array.isArray(value) ? [...new Set(value.filter((v): v is string => typeof v === "string" && v.length > 0))].slice(0, CONFIRMATION_LIMIT) : [];
  } catch {
    return [];
  }
}

// ---- Kimlik eşleşmeleri ----

export interface ConfirmationView {
  readonly key: string;
  readonly before: string;
  readonly literal: string;
  readonly after: string;
  readonly personKey: string;
}

export function listConfirmations(map: Confirmations): ConfirmationView[] {
  return Object.entries(map).map(([key, person]) => {
    const [before = "", literal = "", after = ""] = key.split("|");
    return { key, before, literal, after, personKey: person };
  }).sort((a, b) => a.literal.localeCompare(b.literal, "tr"));
}
