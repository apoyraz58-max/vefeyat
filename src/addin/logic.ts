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
  singleSafe,
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
}

/**
 * Candidates for a selected name, as the VBA "Ölüm tarihi ekle" command
 * gathers them: name matches only (no description matches), one entry per
 * person, strong matches first.
 */
export function lookup(catalog: Catalog, alias: string): Lookup {
  const found = confirmationCandidates(catalog.find(alias, false), alias);
  const { exact, partial } = nameCandidates(found, alias);
  const strong = new Set(exact.map(personKey));
  return {
    query: nameQuery(alias),
    candidates: [...exact, ...partial],
    strong,
    scores: new Map(found.map((person) => [personKey(person), strength(person, alias)])),
    automatic: singleSafe(found, alias) ? (found[0] as Person) : null,
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

export type View = "secili" | "dokum" | "teslim" | "katalog" | "yardim";
export type Tool = "tutarlilik" | "rapor" | "dizin" | "temizlik";

export interface Launch {
  readonly view: View;
  readonly action: "insert" | "spelling" | null;
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
  if (action === "insert" || action === "spelling") return { view: "secili", action, tool: null };
  if (tool === "tutarlilik" || tool === "rapor" || tool === "dizin" || tool === "temizlik") {
    return { view: "teslim", action: null, tool };
  }
  if (view === "dokum" || view === "denetim") return { view: "dokum", action: null, tool: null };
  if (view === "teslim" || view === "katalog" || view === "yardim") return { view, action: null, tool: null };
  if (view === "ayarlar") return { view: "katalog", action: null, tool: null };
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

export type Scope = "first" | "all" | "none";
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
  prefix: { original: "Kaynakta olduğu gibi", o: "ö.", ol: "öl.", v: "v." },
  scope: { first: "Yalnız ilk kullanım", all: "Bütün kullanımlar", none: "Yalnız denetle" },
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
  return { allowed: true, checkEarlier: scope === "first" };
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
